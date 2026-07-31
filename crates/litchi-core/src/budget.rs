//! Hierarchical, thread-safe resource budgets.

use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};

use thiserror::Error;

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
    pub fn child(&self, scope: impl Into<Arc<str>>, limits: Limits) -> Self {
        Self {
            node: Arc::new(Node::new(scope.into(), limits, Some(self.node.clone()))),
        }
    }

    /// Reserves outstanding capacity and releases it when the token is dropped.
    pub fn reserve(&self, resource: Resource, amount: u64) -> Result<Reservation, ResourceLimit> {
        let nodes = self.charge(resource, amount)?;
        Ok(Reservation {
            nodes,
            resource,
            amount,
        })
    }

    /// Charges cumulative work that is not released during this budget's life.
    pub fn consume(&self, resource: Resource, amount: u64) -> Result<(), ResourceLimit> {
        self.charge(resource, amount).map(|_| ())
    }

    /// Current local usage for one resource.
    pub fn used(&self, resource: Resource) -> u64 {
        self.node.used[resource.index()].load(Ordering::Acquire)
    }

    /// Current local limit for one resource.
    pub fn limit(&self, resource: Resource) -> u64 {
        self.node.limits.get(resource)
    }

    fn charge(&self, resource: Resource, amount: u64) -> Result<Vec<Arc<Node>>, ResourceLimit> {
        let mut charged = Vec::new();
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
            current = node.parent.clone();
        }
        Ok(charged)
    }
}

/// RAII token for outstanding budget usage.
#[derive(Debug)]
pub struct Reservation {
    nodes: Vec<Arc<Node>>,
    resource: Resource,
    amount: u64,
}

impl Reservation {
    /// Reserved amount.
    pub const fn amount(&self) -> u64 {
        self.amount
    }

    /// Reserved resource kind.
    pub const fn resource(&self) -> Resource {
        self.resource
    }
}

impl Drop for Reservation {
    fn drop(&mut self) {
        release_nodes(&self.nodes, self.resource, self.amount);
    }
}

fn release_nodes(nodes: &[Arc<Node>], resource: Resource, amount: u64) {
    for node in nodes {
        let counter = &node.used[resource.index()];
        let _ = counter.fetch_update(Ordering::AcqRel, Ordering::Acquire, |used| {
            Some(used.saturating_sub(amount))
        });
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

#[cfg(test)]
mod tests {
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
}
