//! Snapshot transaction layers for DOC embedded-object edits.
//!
//! Opening, mutation, and persistence are separate so each operation can
//! validate a candidate snapshot before publishing it atomically.

mod commit;
mod mutate;
mod open;
