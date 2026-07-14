//! Cycle detection and reachability traversals.

use super::ReferenceGraph;

impl ReferenceGraph {
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
    /// Uses recursive DFS with memoization.
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
        use std::collections::HashSet;

        let mut visited = HashSet::new();
        let mut rec_stack = HashSet::new();

        self.has_cycle_dfs(start_id, &mut visited, &mut rec_stack)
    }

    /// Helper for cycle detection using DFS
    ///
    /// Implements the classical DFS-based cycle detection algorithm for
    /// directed graphs. A cycle exists if we encounter a node that's
    /// currently in the recursion stack (back edge).
    fn has_cycle_dfs(
        &self,
        node: u64,
        visited: &mut std::collections::HashSet<u64>,
        rec_stack: &mut std::collections::HashSet<u64>,
    ) -> bool {
        // Mark current node as visited and add to recursion stack
        visited.insert(node);
        rec_stack.insert(node);

        // Check all outgoing edges
        if let Some(neighbors) = self.get_outgoing_refs(node) {
            for &neighbor in neighbors {
                if !visited.contains(&neighbor) {
                    // Recurse on unvisited neighbor
                    if self.has_cycle_dfs(neighbor, visited, rec_stack) {
                        return true;
                    }
                } else if rec_stack.contains(&neighbor) {
                    // Back edge found - cycle detected
                    return true;
                }
            }
        }

        // Remove from recursion stack before returning
        rec_stack.remove(&node);
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
}
