//! Compatibility re-exports for the iWork object reference graph.
//!
//! The implementation lives in `litchi-iwa-graph`, a dependency-free crate
//! that owns the typed graph vocabulary and algorithms. This module preserves
//! the established `litchi_iwa::ref_graph` path while keeping package parsing
//! and document-format concerns out of the graph crate.
//!
//! # Example
//!
//! ```rust,ignore
//! use litchi_iwa::ref_graph::{ObjectId, ReferenceGraph};
//!
//! let mut graph = ReferenceGraph::new();
//! let source = ObjectId::try_from(1).expect("non-null object ID");
//! let target = ObjectId::try_from(2).expect("non-null object ID");
//! graph.add_object_reference(source, target);
//! assert_eq!(
//!     graph.outgoing(source).map(|ids| ids.collect::<Vec<_>>()),
//!     Some(vec![target])
//! );
//! ```

pub use litchi_iwa_graph::{
    ObjectId, ObjectIdError, ObjectIdIter, ReferenceGraph, ReferenceGraphSnapshot,
    ReferenceGraphStats,
};

#[cfg(test)]
#[allow(
    deprecated,
    reason = "This test deliberately covers retained raw-ID compatibility methods."
)]
mod tests {
    use super::{
        ObjectId, ObjectIdError, ObjectIdIter, ReferenceGraph, ReferenceGraphSnapshot,
        ReferenceGraphStats,
    };

    #[test]
    fn compatibility_path_reexports_typed_graph_api() {
        let source = ObjectId::try_from(1).expect("non-null object ID");
        let target = ObjectId::try_from(2).expect("non-null object ID");
        let mut graph = ReferenceGraph::new();

        graph.add_object_reference(source, target);

        graph.add_reference(source.get(), target.get());
        graph.add_reference(0, target.get());
        assert_eq!(
            graph.get_outgoing_refs(source.get()),
            Some(vec![target.get()])
        );
        assert_eq!(
            graph.get_incoming_refs(target.get()),
            Some(vec![source.get()])
        );

        let outgoing: Option<ObjectIdIter<'_>> = graph.outgoing(source);
        assert_eq!(outgoing.map(|ids| ids.collect::<Vec<_>>()), Some(vec![target]));

        let snapshot: ReferenceGraphSnapshot = graph.snapshot();
        let stats: ReferenceGraphStats = snapshot.statistics();
        assert_eq!(stats.total_objects, 2);

        let invalid = ObjectId::try_from(0);
        assert!(invalid.is_err());
        let error: ObjectIdError = match invalid {
            Ok(_) => return,
            Err(error) => error,
        };
        assert_eq!(error.to_string(), "object identifier must be non-zero");

        let mut root_graph: crate::ReferenceGraph = ReferenceGraph::new();
        root_graph.add_object_reference(source, target);
        assert_eq!(
            root_graph.outgoing(source).map(|ids| ids.collect::<Vec<_>>()),
            Some(vec![target])
        );
    }
}
