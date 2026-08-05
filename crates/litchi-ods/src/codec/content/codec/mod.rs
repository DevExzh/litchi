//! Streaming content XML codec.
//!
//! The traversal implementation is kept behind this small owner facade so
//! content callers do not depend on its parser state or helper functions.

mod traversal;

pub(crate) use traversal::Parser;
