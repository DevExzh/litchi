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

    /// Opens a rough, chunked lease on one resource of this budget.
    ///
    /// Opening claims nothing. The lease claims up to `chunk` units at a time,
    /// the first time it is charged and whenever it runs out, hands them out
    /// locally and never holds more than `chunk` units; see [`Lease`] for what
    /// stays exact and what becomes rough. A `chunk` of zero behaves as one.
    #[must_use]
    pub fn lease(&self, resource: Resource, chunk: u64) -> Lease {
        Lease {
            node: Arc::clone(&self.node),
            resource,
            chunk: chunk.max(1),
            held: 0,
            consumed: 0,
        }
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

/// Claims between `need` and `want` units against `leaf` and then each
/// ancestor in turn, and returns the amount every level now holds.
///
/// Each level is one atomic check-and-add of the current grant, shrunk to the
/// room that level has left. A level with less room than `need` refuses: the
/// levels already charged are released before the refusal is returned, and
/// the refusal names that level, its limit and `used + need`. For a lease
/// whose unspent claim is `amount - need`, that is exactly the value exact
/// accounting of the whole `amount` would have reported there. When a level
/// shrinks the grant, the levels already charged give back the difference, so
/// on success every level holds exactly the returned amount. No level is ever
/// observed above its limit, and the walk takes no reference counts.
fn claim_chain(
    leaf: &Node,
    resource: Resource,
    need: u64,
    want: u64,
) -> Result<u64, ResourceLimit> {
    let mut grant = want.max(need);
    let mut charged = 0usize;
    let mut current = Some(leaf);
    while let Some(node) = current {
        let counter = &node.used[resource.index()];
        let limit = node.limits.get(resource);
        let mut taken = grant;
        let result = counter.fetch_update(Ordering::AcqRel, Ordering::Acquire, |used| {
            let room = limit.checked_sub(used)?;
            if room < need {
                return None;
            }
            taken = grant.min(room);
            used.checked_add(taken)
        });
        match result {
            Ok(_) => {
                if taken < grant {
                    release_ancestor_prefix(leaf, charged, resource, grant - taken);
                    grant = taken;
                }
                charged = charged.saturating_add(1);
            },
            Err(used) => {
                release_ancestor_prefix(leaf, charged, resource, grant);
                return Err(ResourceLimit {
                    resource,
                    observed: used.saturating_add(need),
                    limit,
                    scope: node.scope.clone(),
                });
            },
        }
        current = node.parent.as_deref();
    }
    Ok(grant)
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

/// A rough, chunked claim on one resource of a [`Budget`], made by
/// [`Budget::lease`].
///
/// A lease claims budget a chunk at a time and hands it out locally, so an
/// operation that charges many small amounts pays one atomic update per
/// hierarchy level per chunk instead of one per charge. It never holds more
/// than one chunk: a claim near a limit shrinks, and a refund that would take
/// it past one chunk returns the excess to the budget at once. Every claim is
/// the check-and-add [`Budget::consume`] makes: no level is ever observed
/// above its limit, even transiently, and a refused claim leaves every level
/// as it found it.
///
/// What stays exact: when every charge its holder makes of the resource goes
/// through this one lease, the holder is refused on the same charge, at the
/// same level and with the same [`ResourceLimit`] values as exact accounting
/// would give it, because a claim near a limit shrinks to the room left
/// (never below what the charge needs) and a refusal reports the value the
/// level would have reached had the charge been exact. A charge the holder
/// makes outside the lease, directly or through another lease, sees this
/// lease's unspent units as used. Releasing or dropping the lease returns
/// every unit it holds but has not handed out, so once it is released every
/// counter shows exactly what was handed out.
///
/// What becomes rough: units the lease holds but has not handed out count as
/// used for every other holder of the budget and its ancestors
/// ("pre-claimed"), so another holder can be refused earlier than exact
/// accounting would refuse it, by up to one chunk per other open lease, and
/// observes that much more usage.
///
/// [`Self::refund`] takes back units whose work did not happen, where exact
/// accounting would have dropped a reservation. Dropping the lease releases
/// on every path, including an error and an unwinding panic.
#[derive(Debug)]
pub struct Lease {
    node: Arc<Node>,
    resource: Resource,
    chunk: u64,
    /// Units claimed at every level of the chain and not handed out.
    held: u64,
    /// Units handed out and not refunded; bounds what [`Self::refund`] takes.
    consumed: u64,
}

impl Lease {
    /// Hands out `amount` units, claiming more from the budget first when
    /// the lease holds fewer.
    ///
    /// A claim asks for `max(chunk, amount - held)` units and takes, at each
    /// level, what that level has room for, but never less than the charge
    /// needs. A charge the lease can cover touches no shared state.
    ///
    /// # Errors
    ///
    /// Returns `ResourceLimit` when this budget or an ancestor has less room
    /// than the charge needs. Nothing is handed out or claimed.
    pub fn consume(&mut self, amount: u64) -> Result<(), ResourceLimit> {
        if amount <= self.held {
            self.held -= amount;
        } else {
            let need = amount - self.held;
            let granted = claim_chain(&self.node, self.resource, need, need.max(self.chunk))?;
            // The claim is at least what the charge needed; the rest stays
            // with the lease.
            self.held = granted.saturating_sub(need);
        }
        self.consumed = self.consumed.saturating_add(amount);
        Ok(())
    }

    /// Takes back `amount` units handed out earlier whose work did not
    /// happen.
    ///
    /// The units return to the lease, up to one chunk: whatever the lease
    /// would then hold beyond its chunk goes back to this budget and every
    /// ancestor at once, so a refund never leaves more than one chunk
    /// pre-claimed. Releasing or dropping the lease returns the rest.
    ///
    /// Returns `false`, and changes nothing, when `amount` exceeds what the
    /// lease has handed out and not yet taken back.
    #[must_use = "check whether the refund was accepted"]
    pub fn refund(&mut self, amount: u64) -> bool {
        if amount > self.consumed {
            return false;
        }
        let Some(held) = self.held.checked_add(amount) else {
            return false;
        };
        self.consumed -= amount;
        // A refunded charge can be many chunks wide; keeping all of it would
        // leave it pre-claimed until the lease is released.
        let excess = held.saturating_sub(self.chunk);
        if excess != 0 {
            release_chain(&self.node, self.resource, excess);
        }
        self.held = held - excess;
        true
    }

    /// Returns every unit the lease holds but has not handed out to this
    /// budget and each ancestor. The lease stays usable; its next charge
    /// claims again.
    pub fn release(&mut self) {
        if self.held != 0 {
            release_chain(&self.node, self.resource, self.held);
            self.held = 0;
        }
    }

    /// Units claimed from the budget and not handed out yet.
    #[must_use]
    pub const fn held(&self) -> u64 {
        self.held
    }

    /// Units handed out and not refunded.
    #[must_use]
    pub const fn consumed(&self) -> u64 {
        self.consumed
    }

    /// Units a claim asks for when the lease runs out.
    #[must_use]
    pub const fn chunk(&self) -> u64 {
        self.chunk
    }

    /// Leased resource kind.
    #[must_use]
    pub const fn resource(&self) -> Resource {
        self.resource
    }
}

impl Drop for Lease {
    fn drop(&mut self) {
        self.release();
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
    // -----------------------------------------------------------------------
    // Change 0763: rough budget leases (owner decision 6 of change 0758).
    // -----------------------------------------------------------------------

    /// Work limits of a leaf, its parent and the root.
    fn work_levels(leaf: u64, middle: u64, root: u64) -> [Budget; 3] {
        let work = |limit| Limits::new(100, 100, 100, 100, 100, limit);
        let root = Budget::root("root", work(root));
        let middle = root.child("middle", work(middle));
        let leaf = middle.child("leaf", work(leaf));
        [root, middle, leaf]
    }

    fn work_used(levels: &[Budget; 3]) -> [u64; 3] {
        [
            levels[0].used(Resource::Work),
            levels[1].used(Resource::Work),
            levels[2].used(Resource::Work),
        ]
    }

    /// The first refused charge (its index and refusal) of `amounts` charged
    /// exactly through `Budget::consume`, and every level's usage afterwards.
    fn exact_outcome(
        levels: &[Budget; 3],
        amounts: &[u64],
    ) -> (Option<(usize, ResourceLimit)>, [u64; 3]) {
        for (index, &amount) in amounts.iter().enumerate() {
            if let Err(error) = levels[2].consume(Resource::Work, amount) {
                return (Some((index, error)), work_used(levels));
            }
        }
        (None, work_used(levels))
    }

    /// The same charges through one lease, released at the refusal or the
    /// end, as a writer releases its leases when it is poisoned or finishes.
    fn lease_outcome(
        levels: &[Budget; 3],
        amounts: &[u64],
        chunk: u64,
    ) -> (Option<(usize, ResourceLimit)>, [u64; 3]) {
        let mut lease = levels[2].lease(Resource::Work, chunk);
        for (index, &amount) in amounts.iter().enumerate() {
            let before = lease.held();
            if let Err(error) = lease.consume(amount) {
                // A refused charge claims and hands out nothing.
                assert_eq!(lease.held(), before);
                lease.release();
                return (Some((index, error)), work_used(levels));
            }
            // What the lease holds is claimed at every level, and no level
            // is ever over its limit.
            let used = work_used(levels);
            assert!(used[0] == used[1] && used[1] == used[2]);
            for (level, &limit) in levels.iter().zip(&used) {
                assert!(limit <= level.limit(Resource::Work));
            }
        }
        lease.release();
        (None, work_used(levels))
    }

    #[test]
    fn a_lease_claims_chunks_and_hands_units_out_locally() {
        let [root, middle, leaf] = work_levels(1_000, 1_000, 1_000);
        let mut lease = leaf.lease(Resource::Work, 100);
        assert_eq!(
            (lease.held(), lease.chunk(), lease.resource()),
            (0, 100, Resource::Work)
        );
        assert_eq!(root.used(Resource::Work), 0);

        lease.consume(1).expect("first charge claims a chunk");
        assert_eq!(lease.held(), 99);
        for budget in [&root, &middle, &leaf] {
            assert_eq!(budget.used(Resource::Work), 100);
        }
        lease.consume(99).expect("the rest of the chunk is local");
        assert_eq!(lease.held(), 0);
        assert_eq!(root.used(Resource::Work), 100);
        // A charge larger than a chunk claims what it needs.
        lease.consume(250).expect("a large charge");
        assert_eq!(lease.held(), 0);
        assert_eq!(root.used(Resource::Work), 350);
        lease.consume(0).expect("a zero charge");
        assert_eq!(root.used(Resource::Work), 350);
        lease.consume(1).expect("claims again");
        assert_eq!((lease.held(), lease.consumed()), (99, 351));
        assert_eq!(leaf.used(Resource::Work), 450);

        lease.release();
        for budget in [&root, &middle, &leaf] {
            assert_eq!(budget.used(Resource::Work), 351);
        }
        // A released lease stays usable, and a zero chunk claims one unit.
        let mut exact = leaf.lease(Resource::Work, 0);
        assert_eq!(exact.chunk(), 1);
        exact.consume(3).expect("exact claim");
        assert_eq!((exact.held(), leaf.used(Resource::Work)), (0, 354));
    }

    #[test]
    fn a_sole_lease_holder_is_refused_exactly_as_exact_accounting_refuses() {
        let amounts: Vec<u64> = (0..40_u64).map(|index| 1 + (index * 7) % 9).collect();
        let total: u64 = amounts.iter().sum();
        let loose = total + 100;
        let mut refusals = 0;
        for tight in 0..=total + 2 {
            // The tight limit sits at each level in turn, and at two at once.
            for (leaf, middle, root) in [
                (tight, loose, loose),
                (loose, tight, loose),
                (loose, loose, tight),
                (tight + 3, tight, loose),
            ] {
                let exact = exact_outcome(&work_levels(leaf, middle, root), &amounts);
                for chunk in [1, 2, 5, 16, 64, 1_000, u64::MAX] {
                    let leased = lease_outcome(&work_levels(leaf, middle, root), &amounts, chunk);
                    assert_eq!(
                        leased, exact,
                        "limits {leaf}/{middle}/{root}, chunk {chunk}"
                    );
                }
                refusals += usize::from(exact.0.is_some());
            }
        }
        assert!(refusals > 0);
    }

    #[test]
    fn siblings_see_pre_claimed_units_and_no_level_passes_its_limit() {
        let root = Budget::root("root", Limits::new(100, 100, 100, 100, 100, 100));
        let left = root.child("left", Limits::new(100, 100, 100, 100, 100, 100));
        let right = root.child("right", Limits::new(100, 100, 100, 100, 100, 100));
        let mut lease = left.lease(Resource::Work, 40);
        lease.consume(1).expect("claim");
        assert_eq!(root.used(Resource::Work), 40);
        assert_eq!(left.used(Resource::Work), 40);

        // Exact accounting would admit 1 + 61; the sibling sees the claim.
        let refused = right
            .consume(Resource::Work, 61)
            .expect_err("pre-claimed units count as used");
        assert_eq!(
            (refused.observed, refused.limit, &*refused.scope),
            (101, 100, "root")
        );
        assert_eq!(right.used(Resource::Work), 0);
        right.consume(Resource::Work, 60).expect("up to the limit");
        assert_eq!(root.used(Resource::Work), 100);

        // The holder is still refused only when its own charge cannot fit:
        // it has 39 units left locally and the root is full.
        lease.consume(39).expect("local units");
        let error = lease.consume(1).expect_err("nothing left anywhere");
        assert_eq!(
            (error.observed, error.limit, &*error.scope),
            (101, 100, "root")
        );
        assert_eq!(root.used(Resource::Work), 100);

        drop(lease);
        assert_eq!(root.used(Resource::Work), 100);
        assert_eq!(left.used(Resource::Work), 40);
    }

    #[test]
    fn a_claim_near_a_limit_shrinks_to_the_room_at_every_level() {
        // The root has the least room; the leaf and middle levels must give
        // back what they were charged beyond it.
        let [root, middle, leaf] = work_levels(1_000, 500, 30);
        let mut lease = leaf.lease(Resource::Work, 200);
        lease.consume(5).expect("shrunk claim");
        assert_eq!(lease.held(), 25);
        assert_eq!(
            work_used(&[root.clone(), middle.clone(), leaf.clone()]),
            [30, 30, 30]
        );
        lease.consume(25).expect("the rest of the room");
        let error = lease.consume(1).expect_err("the root is full");
        assert_eq!(
            (error.observed, error.limit, &*error.scope),
            (31, 30, "root")
        );
        assert_eq!(work_used(&[root, middle, leaf]), [30, 30, 30]);
    }

    #[test]
    fn refunds_return_units_to_the_lease_and_release_settles_every_level() {
        let [root, middle, leaf] = work_levels(1_000, 1_000, 1_000);
        let mut lease = leaf.lease(Resource::Work, 64);
        lease.consume(10).expect("charge");
        assert!(lease.refund(4));
        assert_eq!((lease.held(), lease.consumed()), (58, 6));
        assert!(!lease.refund(7), "more than was handed out");
        assert_eq!((lease.held(), lease.consumed()), (58, 6));
        // Refunded units are reused before the budget is touched again.
        lease.consume(58).expect("local");
        assert_eq!(root.used(Resource::Work), 64);
        lease.release();
        for budget in [&root, &middle, &leaf] {
            assert_eq!(budget.used(Resource::Work), 64);
        }
        assert!(lease.refund(64));
        lease.release();
        for budget in [&root, &middle, &leaf] {
            assert_eq!(budget.used(Resource::Work), 0);
        }
    }

    #[test]
    fn a_refund_wider_than_a_chunk_returns_the_excess_at_once() {
        let [root, middle, leaf] = work_levels(100_000, 100_000, 100_000);
        let mut lease = leaf.lease(Resource::Work, 64);
        lease.consume(3).expect("claim a chunk");
        // A charge wider than the chunk uses what the lease holds and claims
        // exactly the rest.
        lease.consume(12_001).expect("wide charge");
        assert_eq!((lease.held(), lease.consumed()), (0, 12_004));
        assert_eq!(root.used(Resource::Work), 12_004);
        // Its work did not happen: the lease keeps one chunk and the budget
        // gets the rest back now, not at release.
        assert!(lease.refund(12_001));
        assert_eq!((lease.held(), lease.consumed()), (64, 3));
        for budget in [&root, &middle, &leaf] {
            assert_eq!(budget.used(Resource::Work), 3 + 64);
        }
        lease.release();
        for budget in [&root, &middle, &leaf] {
            assert_eq!(budget.used(Resource::Work), 3);
        }
    }

    #[test]
    fn a_lease_never_holds_more_than_one_chunk() {
        // Random charges (some wider than a chunk), refunds (some refused),
        // releases and refusals near a three-level limit. After every call
        // the lease holds at most one chunk and every level shows exactly
        // what was handed out plus what the lease holds.
        for (index, chunk) in [1_u64, 3, 64, 4_096, u64::MAX].into_iter().enumerate() {
            let levels = work_levels(50_000, 40_000, 30_000);
            let mut lease = levels[2].lease(Resource::Work, chunk);
            let mut handed_out = 0_u64;
            let (mut refusals, mut refunds) = (0_u32, 0_u32);
            let mut state = 0x9E37_79B9_7F4A_7C15_u64 ^ (u64::try_from(index).expect("index") + 1);
            for _ in 0..20_000 {
                state ^= state << 13;
                state ^= state >> 7;
                state ^= state << 17;
                let roll = state % 100;
                let size = (state >> 16) % 13_000;
                match roll {
                    0..=59 => {
                        let amount = 1 + size % 9;
                        match lease.consume(amount) {
                            Ok(()) => handed_out += amount,
                            Err(error) => {
                                assert_eq!(error.limit, 30_000);
                                refusals += 1;
                            },
                        }
                    },
                    60..=74 => {
                        let amount = 1 + size;
                        match lease.consume(amount) {
                            Ok(()) => handed_out += amount,
                            Err(error) => {
                                assert_eq!(error.observed, handed_out + amount);
                                refusals += 1;
                            },
                        }
                    },
                    75..=94 => {
                        let amount = size.min(lease.consumed() + (state >> 40) % 2);
                        if lease.refund(amount) {
                            handed_out -= amount;
                            refunds += 1;
                        } else {
                            assert!(amount > lease.consumed());
                        }
                    },
                    _ => lease.release(),
                }
                assert!(lease.held() <= lease.chunk(), "chunk {chunk}");
                assert_eq!(lease.consumed(), handed_out);
                for level in &levels {
                    assert_eq!(level.used(Resource::Work), handed_out + lease.held());
                }
            }
            assert!(refusals > 0 && refunds > 0, "chunk {chunk}");
            drop(lease);
            for level in &levels {
                assert_eq!(level.used(Resource::Work), handed_out);
            }
        }
    }

    #[test]
    fn a_lease_releases_when_dropped_even_while_unwinding() {
        let [root, middle, leaf] = work_levels(1_000, 1_000, 1_000);
        let unwound = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            let mut lease = leaf.lease(Resource::Work, 100);
            lease.consume(7).expect("charge");
            assert_eq!(root.used(Resource::Work), 100);
            panic!("abandon the operation");
        }));
        assert!(unwound.is_err());
        for budget in [&root, &middle, &leaf] {
            assert_eq!(budget.used(Resource::Work), 7);
        }
        // An error path drops the lease too.
        let failed = (|| -> Result<(), ResourceLimit> {
            let mut lease = leaf.lease(Resource::Work, 500);
            lease.consume(3)?;
            lease.consume(2_000)?;
            Ok(())
        })();
        assert!(failed.is_err());
        assert_eq!(root.used(Resource::Work), 10);
    }

    #[derive(Debug, Default)]
    struct LeaseWorkerOutcome {
        work_granted: u64,
        work_refusals: u64,
    }

    /// One worker's fixed pseudo-random mix of lease charges, refunds,
    /// releases and re-opened leases with exact consumption and owned and
    /// scoped reservations. Memory is only reserved and released; Work stays
    /// exactly where it was granted and not refunded.
    fn lease_tree_worker(budget: &Budget, chain: &[usize], seed: u64) -> LeaseWorkerOutcome {
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
        let mut outcome = LeaseWorkerOutcome::default();
        let mut owned: Vec<Reservation> = Vec::new();
        let mut lease = budget.lease(Resource::Work, 1 + seed % 40);
        let mut state = seed;
        for _ in 0..TREE_STEPS {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            let amount = 1 + state % 6;
            match (state >> 8) % 8 {
                0..=2 => match lease.consume(amount) {
                    Ok(()) => outcome.work_granted += amount,
                    Err(error) => {
                        refused(&error, Resource::Work);
                        outcome.work_refusals += 1;
                    },
                },
                3 => {
                    let back = amount.min(lease.consumed());
                    assert!(lease.refund(back));
                    outcome.work_granted -= back;
                },
                4 => lease.release(),
                5 => {
                    // Dropping the lease releases it; the next one claims anew.
                    lease = budget.lease(Resource::Work, 1 + (state >> 20) % 40);
                },
                6 => match budget.consume(Resource::Work, amount) {
                    Ok(()) => outcome.work_granted += amount,
                    Err(error) => {
                        refused(&error, Resource::Work);
                        outcome.work_refusals += 1;
                    },
                },
                _ => match budget.reserve(Resource::Memory, amount) {
                    Ok(reservation) if owned.len() < 3 => owned.push(reservation),
                    Ok(reservation) => drop(reservation),
                    Err(error) => refused(&error, Resource::Memory),
                },
            }
            if owned.len() == 3 && (state >> 24) % 2 == 0 {
                owned.clear();
            }
            assert!(lease.held() <= lease.chunk());
        }
        drop(owned);
        drop(lease);
        outcome
    }

    #[test]
    fn concurrent_leases_and_reservations_keep_a_shared_hierarchy_exact() {
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
        let outcomes: Vec<LeaseWorkerOutcome> = std::thread::scope(|scope| {
            // Pre-claimed or not, no level is ever read above its limit.
            let monitor = scope.spawn(|| {
                let mut passes = 0_u64;
                loop {
                    let last = stop.load(Ordering::Acquire);
                    for (index, budget) in budgets.iter().enumerate() {
                        assert!(budget.used(Resource::Memory) <= TREE_MEMORY[index]);
                        assert!(budget.used(Resource::Work) <= TREE_WORK[index]);
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
                    let seed = 0x2545_F491_4F6C_DD1D_u64
                        ^ (u64::try_from(worker).expect("worker index") + 1)
                            .wrapping_mul(0x9E37_79B9_7F4A_7C15);
                    scope.spawn(move || lease_tree_worker(&budget, &tree_chain(node), seed))
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

        // Every lease is gone: each level holds exactly what its subtree was
        // granted and did not refund, and no Memory.
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
        assert!(outcomes.iter().any(|outcome| outcome.work_refusals > 0));
    }
}
