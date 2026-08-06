//! Semantic DOC story and table assembly facade.
//!
//! Each child owns one semantic concern while the inherent `Writer` methods
//! remain available through the same package-level API used by stream assembly.

mod bookmarks;
mod revisions;
mod smart_tags;
mod stories;
mod tables;
