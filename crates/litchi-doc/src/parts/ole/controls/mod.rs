//! Inert OLE-control metadata from the Word binary document tables.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{Control, Controls, Document, Flags, Metadata};
