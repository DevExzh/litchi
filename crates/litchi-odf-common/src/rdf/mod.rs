//! Inert, bounded `RDF`/`XML` metadata for `OpenDocument` packages.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{Graph, Object, Subject, Triple};

pub use package::{
    add_graph, add_triple, graphs, move_triple, remove_graph, remove_triple, replace_graph,
    replace_triple,
};
