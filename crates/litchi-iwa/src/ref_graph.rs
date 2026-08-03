//! Object Reference Graph for iWork Documents
//!
//! Provides a bidirectional graph structure for tracking dependencies between
//! objects in iWork documents. Objects reference each other extensively through
//! TSP.Reference fields, creating a complex object graph.
//!
//! # Features
//!
//! - **Bidirectional tracking**: Both incoming and outgoing references
//! - **Cycle detection**: DFS-based algorithm to detect circular references
//! - **Transitive closure**: BFS to find all reachable objects
//! - **Efficient deduplication**: Inline deduplication using Vec for cache locality
//! - **Memory optimized**: Compact representation for typical small edge lists
//!
//! # Example
//!
//! ```rust,ignore
//! use litchi_iwa::ref_graph::ReferenceGraph;
//!
//! let mut graph = ReferenceGraph::new();
//!
//! // Build the graph
//! graph.add_reference(1, 2);  // Object 1 references object 2
//! graph.add_reference(1, 3);  // Object 1 references object 3
//! graph.add_reference(2, 3);  // Object 2 references object 3
//!
//! // Query dependencies
//! assert_eq!(graph.get_outgoing_refs(1), Some(&vec![2, 3]));
//! assert_eq!(graph.get_incoming_refs(3), Some(&vec![1, 2]));
//!
//! // Check for cycles
//! assert!(!graph.has_cycle_from(1));
//!
//! // Get transitive dependencies
//! let reachable = graph.get_reachable(1);
//! assert_eq!(reachable.len(), 3);
//! ```

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

/// A validated, non-null iWork object identifier.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct ObjectId(NonZeroU64);

impl ObjectId {
    /// Construct an object identifier, rejecting the protobuf null sentinel.
    pub const fn new(raw: u64) -> Option<Self> {
        match NonZeroU64::new(raw) {
            Some(value) => Some(Self(value)),
            None => None,
        }
    }

    /// Return the native identifier used at the wire boundary.
    pub const fn get(self) -> u64 {
        self.0.get()
    }
}

impl TryFrom<u64> for ObjectId {
    type Error = ObjectIdError;

    fn try_from(raw: u64) -> Result<Self, Self::Error> {
        Self::new(raw).ok_or(ObjectIdError)
    }
}

/// Error returned when a null protobuf reference is used as an object ID.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ObjectIdError;

impl fmt::Display for ObjectIdError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("object identifier must be non-zero")
    }
}

impl std::error::Error for ObjectIdError {}

#[derive(Debug, Clone, Default)]
struct GraphState {
    incoming_refs: HashMap<u64, Vec<u64>>,
    outgoing_refs: HashMap<u64, Vec<u64>>,
}

/// Object reference graph for tracking dependencies
///
/// Maintains a bidirectional graph of object references with efficient
/// deduplication and cache-friendly memory layout.
///
/// # Memory Optimization
///
/// - Uses `Vec` instead of `HashSet` for edge lists (typically small, <10 edges)
/// - Deduplicates inline to avoid redundant storage
/// - Compact representation reduces cache misses during traversal
///
/// # Performance Characteristics
///
/// - Add reference: O(n) where n is edge count per node (typically <10)
/// - Get references: O(1) HashMap lookup
/// - Cycle detection: O(V + E) where V is vertices, E is edges
/// - Transitive closure: O(V + E) BFS traversal
#[derive(Debug, Clone)]
pub struct ReferenceGraph {
    state: Arc<GraphState>,
}

/// An immutable, cheaply shareable typed view of a reference graph.
#[derive(Debug, Clone)]
pub struct ReferenceGraphSnapshot {
    state: Arc<GraphState>,
}

/// Allocation-free iterator over validated object IDs in one edge list.
#[derive(Debug, Clone)]
pub struct ObjectIdIter<'a> {
    inner: std::slice::Iter<'a, u64>,
}

impl Iterator for ObjectIdIter<'_> {
    type Item = ObjectId;

    fn next(&mut self) -> Option<Self::Item> {
        self.inner.find_map(|&raw| ObjectId::new(raw))
    }
}

impl Default for ReferenceGraph {
    fn default() -> Self {
        Self::new()
    }
}

impl ReferenceGraph {
    /// Create an empty reference graph
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let graph = ReferenceGraph::new();
    /// assert!(graph.is_empty());
    /// ```
    pub fn new() -> Self {
        Self {
            state: Arc::new(GraphState::default()),
        }
    }

    /// Freeze the current graph into a cheap immutable snapshot.
    pub fn snapshot(&self) -> ReferenceGraphSnapshot {
        ReferenceGraphSnapshot {
            state: Arc::clone(&self.state),
        }
    }

    /// Add a reference through the checked typed identity path.
    pub fn add_object_reference(&mut self, source: ObjectId, target: ObjectId) {
        self.add_reference(source.get(), target.get());
    }

    /// Add a reference from source to target
    ///
    /// Automatically deduplicates to avoid storing the same edge multiple times.
    /// This is important because protobuf messages may contain duplicate references.
    ///
    /// # Arguments
    ///
    /// * `source_id` - The object that contains the reference
    /// * `target_id` - The object being referenced
    ///
    /// # Performance
    ///
    /// O(n) where n is the number of existing edges (typically <10).
    /// Linear search is faster than HashSet for small n due to better cache locality.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let mut graph = ReferenceGraph::new();
    /// graph.add_reference(1, 2);
    /// graph.add_reference(1, 2);  // Duplicate - will be ignored
    /// assert_eq!(graph.get_outgoing_refs(1), Some(&vec![2]));
    /// ```
    pub fn add_reference(&mut self, source_id: u64, target_id: u64) {
        let state = Arc::make_mut(&mut self.state);
        // Add to outgoing refs with deduplication
        let outgoing = state.outgoing_refs.entry(source_id).or_default();
        if !outgoing.contains(&target_id) {
            outgoing.push(target_id);
        }

        // Add to incoming refs with deduplication
        let incoming = state.incoming_refs.entry(target_id).or_default();
        if !incoming.contains(&source_id) {
            incoming.push(source_id);
        }
    }

    /// Get objects that reference the given object (incoming edges)
    ///
    /// Returns the "dependents" - objects that point to this one.
    ///
    /// # Arguments
    ///
    /// * `object_id` - The target object ID
    ///
    /// # Returns
    ///
    /// Optional reference to a vector of referencing object IDs
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let mut graph = ReferenceGraph::new();
    /// graph.add_reference(1, 2);
    /// graph.add_reference(3, 2);
    /// assert_eq!(graph.get_incoming_refs(2), Some(&vec![1, 3]));
    /// ```
    pub fn get_incoming_refs(&self, object_id: u64) -> Option<&Vec<u64>> {
        self.state.incoming_refs.get(&object_id)
    }

    /// Get objects referenced by the given object (outgoing edges)
    ///
    /// Returns the "dependencies" - objects this one points to.
    ///
    /// # Arguments
    ///
    /// * `object_id` - The source object ID
    ///
    /// # Returns
    ///
    /// Optional reference to a vector of referenced object IDs
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let mut graph = ReferenceGraph::new();
    /// graph.add_reference(1, 2);
    /// graph.add_reference(1, 3);
    /// assert_eq!(graph.get_outgoing_refs(1), Some(&vec![2, 3]));
    /// ```
    pub fn get_outgoing_refs(&self, object_id: u64) -> Option<&Vec<u64>> {
        self.state.outgoing_refs.get(&object_id)
    }

    /// Get all object IDs in the graph
    ///
    /// Returns a set containing all objects that either have outgoing references
    /// or are referenced by other objects.
    ///
    /// # Performance
    ///
    /// O(V) where V is the number of unique objects
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let mut graph = ReferenceGraph::new();
    /// graph.add_reference(1, 2);
    /// graph.add_reference(2, 3);
    /// let all = graph.all_objects();
    /// assert_eq!(all.len(), 3);
    /// ```
    pub fn all_objects(&self) -> std::collections::HashSet<u64> {
        let mut all = std::collections::HashSet::new();
        all.extend(self.state.incoming_refs.keys());
        all.extend(self.state.outgoing_refs.keys());
        all
    }

    /// Check if there's a cycle reachable from the given object
    ///
    /// Uses depth-first search with a visited set to detect back edges.
    /// This is useful for validating document integrity and detecting
    /// corrupted or malformed iWork files.
    ///
    /// # Arguments
    ///
    /// * `start_id` - The object ID to start checking from
    ///
    /// # Returns
    ///
    /// `true` if a cycle is detected, `false` otherwise
    ///
    /// # Performance
    ///
    /// O(V + E) where V is vertices and E is edges in the reachable subgraph.
    /// Uses an explicit stack so malformed deep graphs cannot overflow the
    /// call stack.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let mut graph = ReferenceGraph::new();
    /// graph.add_reference(1, 2);
    /// graph.add_reference(2, 3);
    /// graph.add_reference(3, 1);  // Creates cycle
    /// assert!(graph.has_cycle_from(1));
    /// ```
    pub fn has_cycle_from(&self, start_id: u64) -> bool {
        // 0 = unseen, 1 = active, 2 = complete. Each node is pushed once to
        // enter and once to leave, replacing recursive DFS with bounded stack
        // growth in the heap.
        let mut colors = HashMap::new();
        let mut stack = vec![(start_id, false)];

        while let Some((node, leaving)) = stack.pop() {
            if leaving {
                colors.insert(node, 2_u8);
                continue;
            }

            match colors.get(&node).copied() {
                Some(1) => return true,
                Some(2) => continue,
                _ => {},
            }

            colors.insert(node, 1);
            stack.push((node, true));

            if let Some(neighbors) = self.get_outgoing_refs(node) {
                for &neighbor in neighbors.iter().rev() {
                    match colors.get(&neighbor).copied() {
                        Some(1) => return true,
                        Some(2) => continue,
                        _ => stack.push((neighbor, false)),
                    }
                }
            }
        }

        false
    }

    /// Get all objects reachable from the given object via BFS
    ///
    /// Performs breadth-first traversal to find all transitively referenced objects.
    /// Useful for:
    /// - Extracting complete sub-documents
    /// - Determining what needs to be loaded to fully resolve an object
    /// - Computing dependency closures
    ///
    /// # Arguments
    ///
    /// * `start_id` - The starting object ID
    ///
    /// # Returns
    ///
    /// Vector of all reachable object IDs (including `start_id`)
    ///
    /// # Performance
    ///
    /// O(V + E) where V is vertices and E is edges in the reachable subgraph.
    /// Uses BFS for cache-friendly traversal (better locality than DFS).
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let mut graph = ReferenceGraph::new();
    /// graph.add_reference(1, 2);
    /// graph.add_reference(1, 3);
    /// graph.add_reference(2, 4);
    /// let reachable = graph.get_reachable(1);
    /// assert_eq!(reachable.len(), 4);  // [1, 2, 3, 4]
    /// ```
    pub fn get_reachable(&self, start_id: u64) -> Vec<u64> {
        use std::collections::{HashSet, VecDeque};

        let mut visited = HashSet::new();
        let mut queue = VecDeque::new();
        let mut result = Vec::new();

        // Start with the initial object
        queue.push_back(start_id);
        visited.insert(start_id);

        // BFS traversal
        while let Some(node) = queue.pop_front() {
            result.push(node);

            // Add all unvisited neighbors to the queue
            if let Some(neighbors) = self.get_outgoing_refs(node) {
                for &neighbor in neighbors {
                    if visited.insert(neighbor) {
                        queue.push_back(neighbor);
                    }
                }
            }
        }

        result
    }

    /// Get the number of objects in the graph
    ///
    /// Returns the count of unique objects that participate in the reference graph.
    ///
    /// # Performance
    ///
    /// O(V) where V is the number of unique objects
    #[inline]
    pub fn len(&self) -> usize {
        self.all_objects().len()
    }

    /// Check if the graph is empty
    ///
    /// Returns `true` if there are no references in the graph.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.state.incoming_refs.is_empty() && self.state.outgoing_refs.is_empty()
    }

    /// Get the total number of edges (references) in the graph
    ///
    /// Returns the total count of references between objects.
    ///
    /// # Performance
    ///
    /// O(V) where V is the number of objects with outgoing references
    #[inline]
    pub fn edge_count(&self) -> usize {
        self.state.outgoing_refs.values().map(|v| v.len()).sum()
    }

    /// Get statistics about the reference graph
    ///
    /// Returns a tuple of:
    /// - `total_objects`: Number of unique objects in the graph
    /// - `total_edges`: Total number of references
    /// - `max_out_degree`: Maximum number of outgoing references from any object
    /// - `max_in_degree`: Maximum number of incoming references to any object
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let (objects, edges, max_out, max_in) = graph.stats();
    /// println!("Graph: {} objects, {} edges", objects, edges);
    /// println!("Max out-degree: {}, max in-degree: {}", max_out, max_in);
    /// ```
    pub fn stats(&self) -> (usize, usize, usize, usize) {
        let total_objects = self.len();
        let total_edges: usize = self.state.outgoing_refs.values().map(|v| v.len()).sum();
        let max_out_degree = self
            .state
            .outgoing_refs
            .values()
            .map(|v| v.len())
            .max()
            .unwrap_or(0);
        let max_in_degree = self
            .state
            .incoming_refs
            .values()
            .map(|v| v.len())
            .max()
            .unwrap_or(0);

        (total_objects, total_edges, max_out_degree, max_in_degree)
    }
}

impl ReferenceGraphSnapshot {
    /// Get the validated objects that reference `object_id`.
    ///
    /// The iterator borrows the immutable snapshot and performs no allocation.
    pub fn incoming(&self, object_id: ObjectId) -> Option<ObjectIdIter<'_>> {
        self.state
            .incoming_refs
            .get(&object_id.get())
            .map(|ids| ObjectIdIter { inner: ids.iter() })
    }

    /// Get the validated objects referenced by `object_id`.
    ///
    /// The iterator borrows the immutable snapshot and performs no allocation.
    pub fn outgoing(&self, object_id: ObjectId) -> Option<ObjectIdIter<'_>> {
        self.state
            .outgoing_refs
            .get(&object_id.get())
            .map(|ids| ObjectIdIter { inner: ids.iter() })
    }

    /// Return every non-null object participating in the graph.
    pub fn all_object_ids(&self) -> HashSet<ObjectId> {
        self.state
            .incoming_refs
            .keys()
            .chain(self.state.outgoing_refs.keys())
            .filter_map(|&raw| ObjectId::new(raw))
            .collect()
    }

    /// Check whether a cycle is reachable from `object_id`.
    pub fn has_cycle_from(&self, object_id: ObjectId) -> bool {
        ReferenceGraph {
            state: Arc::clone(&self.state),
        }
        .has_cycle_from(object_id.get())
    }

    /// Return the non-null objects reachable from `object_id`, including it.
    pub fn reachable(&self, object_id: ObjectId) -> Vec<ObjectId> {
        ReferenceGraph {
            state: Arc::clone(&self.state),
        }
        .get_reachable(object_id.get())
        .into_iter()
        .filter_map(ObjectId::new)
        .collect()
    }

    /// Return the number of non-null objects in the snapshot.
    #[inline]
    pub fn len(&self) -> usize {
        self.all_object_ids().len()
    }

    /// Return whether the snapshot contains no non-null object references.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.all_object_ids().is_empty()
    }

    /// Return the number of stored non-duplicate edges with non-null endpoints.
    #[inline]
    pub fn edge_count(&self) -> usize {
        self.state
            .outgoing_refs
            .iter()
            .filter_map(|(&source, targets)| ObjectId::new(source).map(|_| targets))
            .map(|targets| {
                targets
                    .iter()
                    .filter(|&&target| ObjectId::new(target).is_some())
                    .count()
            })
            .sum()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_reference_graph_basic() {
        let mut graph = ReferenceGraph::new();

        graph.add_reference(1, 2);
        graph.add_reference(1, 3);
        graph.add_reference(2, 3);

        assert_eq!(graph.get_outgoing_refs(1), Some(&vec![2, 3]));
        assert_eq!(graph.get_incoming_refs(3), Some(&vec![1, 2]));
        assert_eq!(graph.get_incoming_refs(1), None);
    }

    #[test]
    fn test_reference_graph_deduplication() {
        let mut graph = ReferenceGraph::new();

        // Add the same reference multiple times
        graph.add_reference(1, 2);
        graph.add_reference(1, 2);
        graph.add_reference(1, 2);

        // Should only appear once
        assert_eq!(graph.get_outgoing_refs(1), Some(&vec![2]));
        assert_eq!(graph.get_incoming_refs(2), Some(&vec![1]));
    }

    #[test]
    fn test_reference_graph_cycle_detection() {
        let mut graph = ReferenceGraph::new();

        // Create a simple cycle: 1 -> 2 -> 3 -> 1
        graph.add_reference(1, 2);
        graph.add_reference(2, 3);
        graph.add_reference(3, 1);

        assert!(graph.has_cycle_from(1));
        assert!(graph.has_cycle_from(2));
        assert!(graph.has_cycle_from(3));
    }

    #[test]
    fn test_reference_graph_no_cycle() {
        let mut graph = ReferenceGraph::new();

        // Create a DAG: 1 -> 2 -> 3
        //                 \-> 4
        graph.add_reference(1, 2);
        graph.add_reference(1, 4);
        graph.add_reference(2, 3);

        assert!(!graph.has_cycle_from(1));
        assert!(!graph.has_cycle_from(2));
        assert!(!graph.has_cycle_from(3));
        assert!(!graph.has_cycle_from(4));
    }

    #[test]
    fn test_reference_graph_reachability() {
        let mut graph = ReferenceGraph::new();

        // Create graph: 1 -> 2 -> 3
        //                \-> 4 -> 5
        graph.add_reference(1, 2);
        graph.add_reference(1, 4);
        graph.add_reference(2, 3);
        graph.add_reference(4, 5);

        let reachable = graph.get_reachable(1);
        assert_eq!(reachable.len(), 5);
        assert!(reachable.contains(&1));
        assert!(reachable.contains(&2));
        assert!(reachable.contains(&3));
        assert!(reachable.contains(&4));
        assert!(reachable.contains(&5));

        let reachable_from_2 = graph.get_reachable(2);
        assert_eq!(reachable_from_2.len(), 2);
        assert!(reachable_from_2.contains(&2));
        assert!(reachable_from_2.contains(&3));
    }

    #[test]
    fn test_reference_graph_stats() {
        let mut graph = ReferenceGraph::new();

        graph.add_reference(1, 2);
        graph.add_reference(1, 3);
        graph.add_reference(1, 4);
        graph.add_reference(2, 3);

        let (objects, edges, max_out, max_in) = graph.stats();
        assert_eq!(objects, 4);
        assert_eq!(edges, 4);
        assert_eq!(max_out, 3); // Node 1 has 3 outgoing edges
        assert_eq!(max_in, 2); // Node 3 has 2 incoming edges
    }

    #[test]
    fn test_reference_graph_empty() {
        let graph = ReferenceGraph::new();

        assert!(graph.is_empty());
        assert_eq!(graph.len(), 0);

        let (objects, edges, max_out, max_in) = graph.stats();
        assert_eq!(objects, 0);
        assert_eq!(edges, 0);
        assert_eq!(max_out, 0);
        assert_eq!(max_in, 0);
    }

    #[test]
    fn test_all_objects() {
        let mut graph = ReferenceGraph::new();

        graph.add_reference(1, 2);
        graph.add_reference(3, 4);
        graph.add_reference(5, 6);

        let all = graph.all_objects();
        assert_eq!(all.len(), 6);
        for i in 1..=6 {
            assert!(all.contains(&i));
        }
    }

    #[test]
    fn object_ids_reject_null_and_round_trip() {
        assert_eq!(ObjectId::new(0), None);

        let object_id = ObjectId::try_from(42).expect("non-zero IDs are valid");
        assert_eq!(object_id.get(), 42);
        assert!(ObjectId::try_from(0).is_err());
        assert_eq!(std::mem::size_of::<ObjectId>(), std::mem::size_of::<u64>());
    }

    #[test]
    fn reference_graph_values_are_send_and_sync() {
        fn assert_send_sync<T: Send + Sync>() {}

        assert_send_sync::<ReferenceGraph>();
        assert_send_sync::<ReferenceGraphSnapshot>();
    }

    #[test]
    fn typed_snapshot_is_immutable_and_filters_null_edges() {
        let mut graph = ReferenceGraph::new();
        let one = ObjectId::try_from(1).unwrap();
        let two = ObjectId::try_from(2).unwrap();
        let three = ObjectId::try_from(3).unwrap();

        graph.add_object_reference(one, two);
        // Keep the legacy wire-facing path covered: null protobuf references
        // are tolerated at the boundary but never exposed by the typed view.
        graph.add_reference(one.get(), 0);
        let snapshot = graph.snapshot();

        graph.add_object_reference(one, three);

        assert_eq!(
            snapshot.outgoing(one).unwrap().collect::<Vec<_>>(),
            vec![two]
        );
        assert_eq!(
            snapshot.incoming(two).unwrap().collect::<Vec<_>>(),
            vec![one]
        );
        assert_eq!(snapshot.reachable(one), vec![one, two]);
        assert!(!snapshot.has_cycle_from(one));
        assert_eq!(snapshot.all_object_ids(), HashSet::from([one, two]));
        assert_eq!(snapshot.len(), 2);
        assert_eq!(snapshot.edge_count(), 1);
    }

    #[test]
    fn typed_snapshot_detects_cycles_and_deduplicates_edges() {
        let mut graph = ReferenceGraph::new();
        let one = ObjectId::try_from(1).unwrap();
        let two = ObjectId::try_from(2).unwrap();
        let three = ObjectId::try_from(3).unwrap();

        graph.add_object_reference(one, two);
        graph.add_object_reference(one, two);
        graph.add_object_reference(two, three);
        graph.add_object_reference(three, one);

        let snapshot = graph.snapshot();

        assert_eq!(snapshot.edge_count(), 3);
        assert!(snapshot.has_cycle_from(one));
        assert_eq!(snapshot.reachable(one), vec![one, two, three]);
    }

    #[test]
    fn cycle_detection_handles_deep_graphs_without_recursion() {
        const CHAIN_LENGTH: u64 = 10_000;
        let mut graph = ReferenceGraph::new();

        for object_id in 1..CHAIN_LENGTH {
            graph.add_reference(object_id, object_id + 1);
        }

        assert!(!graph.has_cycle_from(1));
        graph.add_reference(CHAIN_LENGTH, 1);
        assert!(graph.has_cycle_from(1));
    }
}
