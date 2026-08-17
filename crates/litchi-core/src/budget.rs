//! Hierarchical, thread-safe resource budgets.

use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};

use smallvec::SmallVec;
use thiserror::Error;

type ChargedNodes = SmallVec<[Arc<Node>; 4]>;

const RESOURCE_COUNT: usize = 6;

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
            values: [memory, input_bytes, output_bytes, objects, depth, work],
        }
    }

    /// Conservative named defaults. Workload-specific limits remain explicit.
    #[must_use]
    pub const fn for_profile(profile: Profile) -> Self {
        const MIB: u64 = 1024 * 1024;
        const GIB: u64 = 1024 * MIB;
        match profile {
            Profile::Server => {
                Self::new(256 * MIB, 2 * GIB, 4 * GIB, 10_000_000, 256, 1_000_000_000)
            },
            Profile::Desktop => Self::new(GIB, 8 * GIB, 16 * GIB, 50_000_000, 512, 5_000_000_000),
            Profile::TrustedBatch => Self::new(
                4 * GIB,
                64 * GIB,
                128 * GIB,
                250_000_000,
                1024,
                50_000_000_000,
            ),
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
        let nodes = self.charge(resource, amount)?;
        Ok(Reservation {
            nodes,
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
        let mut charged = 0usize;
        let mut current = Some(self.node.as_ref());
        while let Some(node) = current {
            let counter = &node.used[resource.index()];
            let limit = node.limits.get(resource);
            let result = counter.fetch_update(Ordering::AcqRel, Ordering::Acquire, |used| {
                used.checked_add(amount).filter(|next| *next <= limit)
            });
            match result {
                Ok(_) => charged = charged.saturating_add(1),
                Err(used) => {
                    release_ancestor_prefix(self.node.as_ref(), charged, resource, amount);
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

    fn charge(&self, resource: Resource, amount: u64) -> Result<ChargedNodes, ResourceLimit> {
        let mut charged = ChargedNodes::new();
        let mut current = Some(self.node.clone());
        while let Some(node) = current {
            let counter = &node.used[resource.index()];
            let limit = node.limits.get(resource);
            let result = counter.fetch_update(Ordering::AcqRel, Ordering::Acquire, |used| {
                used.checked_add(amount).filter(|next| *next <= limit)
            });
            match result {
                Ok(_) => charged.push(node.clone()),
                Err(used) => {
                    release_nodes(&charged, resource, amount);
                    return Err(ResourceLimit {
                        resource,
                        observed: used.saturating_add(amount),
                        limit,
                        scope: node.scope.clone(),
                    });
                },
            }
            let parent = node.parent.clone();
            current = parent;
        }
        Ok(charged)
    }
}

/// RAII token for outstanding budget usage.
#[derive(Debug)]
pub struct Reservation {
    nodes: ChargedNodes,
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
            release_nodes(&self.nodes, self.resource, self.amount - amount);
        }
        self.nodes.clear();
        self.amount = 0;
        true
    }
}

impl Drop for Reservation {
    fn drop(&mut self) {
        release_nodes(&self.nodes, self.resource, self.amount);
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

fn release_nodes(nodes: &[Arc<Node>], resource: Resource, amount: u64) {
    for node in nodes {
        release_node(node, resource, amount);
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
    fn common_hierarchies_keep_reservation_nodes_inline() {
        let root = Budget::root("root", limits(100));
        let child = root.child("child", limits(100));
        let grandchild = child.child("grandchild", limits(100));
        let leaf = grandchild.child("leaf", limits(100));
        let reservation = leaf
            .reserve(Resource::Memory, 1)
            .expect("four-level reservation");

        assert_eq!(reservation.nodes.len(), 4);
        assert!(!reservation.nodes.spilled());
        drop(reservation);
        assert_eq!(root.used(Resource::Memory), 0);
    }

    #[test]
    fn deep_hierarchies_spill_and_still_roll_back_exactly() {
        let root = Budget::root("root", limits(1));
        let first = root.child("first", limits(100));
        let second = first.child("second", limits(100));
        let third = second.child("third", limits(100));
        let fourth = third.child("fourth", limits(100));
        let leaf = fourth.child("leaf", limits(100));

        let reservation = leaf
            .reserve(Resource::Memory, 1)
            .expect("six-level reservation");
        assert_eq!(reservation.nodes.len(), 6);
        assert!(reservation.nodes.spilled());
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
}
